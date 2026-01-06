/**
 * 图片解析器
 * 处理图片资源的提取和转换
 */

import JSZip from 'jszip';
import { log } from '../utils';
import type { RelsMap, PptxParseResult } from './types';

/**
 * 解析图片资源
 * @param zip JSZip对象
 * @param result 解析结果对象
 */
export async function parseImages(zip: JSZip, result: PptxParseResult): Promise<void> {
  try {
    const imageMap = new Map<string, string>();

    // 遍历所有幻灯片
    for (const slide of result.slides) {
      // 1. 处理背景图片
      if (slide.background && typeof slide.background !== 'string') {
        const bg = slide.background as any;
        if (bg.type === 'image' && bg.relId && !imageMap.has(bg.relId)) {
          // 从relsMap获取目标路径
          const rel = slide.relsMap[bg.relId];
          if (rel) {
            const imagePath = rel.target.startsWith('..')
              ? `ppt/${rel.target.substring(3)}`
              : rel.target;

            const imageFile = await zip.file(imagePath);
            if (imageFile) {
              const base64 = await imageFile.async('base64');
              const mimeType = inferMimeType(imagePath);
              imageMap.set(bg.relId, `data:${mimeType};base64,${base64}`);

              // 更新背景的value为base64 URL
              bg.value = imageMap.get(bg.relId);
            }
          }
        }
      }

      // 2. 处理元素中的图片
      for (const element of slide.elements) {
        // @ts-ignore - 类型扩展
        if (element.type === 'image') {
          const imgElement = element as any;
          const relId = imgElement.relId;

          if (relId && !imageMap.has(relId)) {
            // 从relsMap获取目标路径
            const rel = slide.relsMap[relId];
            if (rel) {
              // 读取图片文件
              const imagePath = rel.target.startsWith('..')
                ? `ppt/${rel.target.substring(3)}`
                : rel.target;

              const imageFile = await zip.file(imagePath);
              if (imageFile) {
                const base64 = await imageFile.async('base64');
                const mimeType = inferMimeType(imagePath);

                imageMap.set(relId, `data:${mimeType};base64,${base64}`);

                // 更新元素的src
                imgElement.src = imageMap.get(relId);
              }
            }
          }
        }
      }
    }

    log('info', `Parsed ${imageMap.size} images`);
  } catch (error) {
    log('warn', 'Failed to parse images', error);
  }
}

/**
 * 从文件路径推断MIME类型
 * @param filePath 文件路径
 * @returns MIME类型
 */
function inferMimeType(filePath: string): string {
  const ext = filePath.split('.').pop()?.toLowerCase();

  switch (ext) {
    case 'png':
      return 'image/png';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    case 'svg':
      return 'image/svg+xml';
    case 'webp':
      return 'image/webp';
    default:
      return 'image/jpeg';
  }
}
