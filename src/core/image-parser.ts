/**
 * 图片解析器
 * 处理图片资源的提取和转换
 */

import JSZip from 'jszip';
import { log } from '../utils/index';
import type { RelsMap, PptxParseResult, SlideLayoutResult, MasterSlideResult } from './types';

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
      // 1. 处理幻灯片背景图片
      if (slide.background && typeof slide.background !== 'string') {
        const bg = slide.background as any;
        if (bg.type === 'image' && bg.relId) {
          const imageUrl = await resolveImageRelId(zip, bg.relId, slide.relsMap, imageMap);
          if (imageUrl) {
            bg.value = imageUrl;
          }
        }
      }

      // 2. 处理幻灯片元素中的图片
      let slideImageCount = 0;
      for (const element of slide.elements) {
        // @ts-ignore - 类型扩展
        if (element.type === 'image') {
          slideImageCount++;
          const imgElement = element as any;
          const relId = imgElement.relId;
          if (relId) {
            log('info', `Found image element in slide: relId=${relId}, src=${imgElement.src}`);
            const imageUrl = await resolveImageRelId(zip, relId, slide.relsMap, imageMap);
            if (imageUrl) {
              imgElement.src = imageUrl;
              log('info', `Updated image src to base64 URL, length=${imageUrl.length}`);
            } else {
              log('warn', `Failed to resolve image relId=${relId}`);
            }
          } else {
            log('warn', 'Image element missing relId');
          }
        }
      }
      if (slideImageCount > 0) {
        log('info', `Found ${slideImageCount} image elements in slide ${slide.id}`);
      }
    }

    // 3. 处理所有 layout 的背景图片和元素图片
    if (result.slideLayouts) {
      const layoutCount = Object.keys(result.slideLayouts).length;
      log('info', `Processing ${layoutCount} slide layouts`);
      for (const layoutId in result.slideLayouts) {
        const layout = result.slideLayouts[layoutId];
        log('info', `Processing layout: ${layoutId}`);
        await processLayoutOrMasterImages(zip, layout, layout.relsMap, imageMap);
      }
    }

    // 4. 处理所有 master 的背景图片和元素图片
    if (result.masterSlides) {
      log('info', `Processing ${result.masterSlides.length} master slides`);
      for (const master of result.masterSlides) {
        if (master.relsMap) {
          log('info', `Processing master: ${master.id}`);
          await processLayoutOrMasterImages(zip, master, master.relsMap, imageMap);
        } else {
          log('warn', `Master ${master.id} missing relsMap`);
        }
      }
    }

    // 保存 mediaMap 到结果中
    result.mediaMap = imageMap;

    log('info', `Parsed ${imageMap.size} images`);
  } catch (error) {
    log('warn', 'Failed to parse images', error);
  }
}

/**
 * 处理 layout 或 master 中的图片
 */
async function processLayoutOrMasterImages(
  zip: JSZip,
  layoutOrMaster: SlideLayoutResult | MasterSlideResult,
  relsMap: RelsMap,
  imageMap: Map<string, string>
): Promise<void> {
  // 处理背景图片
  if (layoutOrMaster.background && typeof layoutOrMaster.background !== 'string') {
    const bg = layoutOrMaster.background as any;
    if (bg.type === 'image' && bg.relId) {
      const imageUrl = await resolveImageRelId(zip, bg.relId, relsMap, imageMap);
      if (imageUrl) {
        bg.value = imageUrl;
      }
    }
  }

  // 处理元素中的图片
  let elementImageCount = 0;
  for (const element of layoutOrMaster.elements) {
    // @ts-ignore - 类型扩展
    if (element.type === 'image') {
      elementImageCount++;
      const imgElement = element as any;
      const relId = imgElement.relId;
      if (relId) {
        log('info', `Found image element in ${layoutOrMaster.id}: relId=${relId}, src=${imgElement.src}`);
        const imageUrl = await resolveImageRelId(zip, relId, relsMap, imageMap);
        if (imageUrl) {
          imgElement.src = imageUrl;
          log('info', `Updated image src to base64 URL, length=${imageUrl.length}`);
        } else {
          log('warn', `Failed to resolve image relId=${relId} in ${layoutOrMaster.id}`);
        }
      } else {
        log('warn', `Image element missing relId in ${layoutOrMaster.id}`);
      }
    }
  }
  if (elementImageCount > 0) {
    log('info', `Found ${elementImageCount} image elements in ${layoutOrMaster.id}`);
  }
}

/**
 * 在ZIP中查找图片文件（支持不区分大小写和多种路径）
 */
async function findImageFileInZip(zip: JSZip, targetPath: string): Promise<{ file: any; path: string } | null> {
  log('info', `findImageFileInZip: targetPath=${targetPath}`);
  // 尝试1: 原始路径（精确匹配）
  let imageFile = await zip.file(targetPath);
  if (imageFile) {
    return { file: imageFile, path: targetPath };
  }

  // 尝试2: 如果以..开头，转换为ppt/路径
  if (targetPath.startsWith('..')) {
    const normalizedPath = `ppt/${targetPath.substring(3)}`;
    imageFile = await zip.file(normalizedPath);
    if (imageFile) {
      return { file: imageFile, path: normalizedPath };
    }
  }

  // 尝试3: 如果包含media/但没有ppt/前缀，添加ppt/前缀
  if (targetPath.includes('media/') && !targetPath.startsWith('ppt/')) {
    const normalizedPath = `ppt/${targetPath}`;
    imageFile = await zip.file(normalizedPath);
    if (imageFile) {
      return { file: imageFile, path: normalizedPath };
    }
  }

  // 尝试4: 如果以media/开头，添加ppt/前缀
  if (targetPath.startsWith('media/')) {
    const normalizedPath = `ppt/${targetPath}`;
    imageFile = await zip.file(normalizedPath);
    if (imageFile) {
      return { file: imageFile, path: normalizedPath };
    }
  }

  // 尝试5: 针对常见PPTX目录的路径尝试
  const commonPrefixes = [
    'ppt/slideLayouts/',
    'ppt/slideMasters/',
    'ppt/slides/',
    'ppt/notesMasters/',
    'ppt/notesSlides/',
    'ppt/',
    ''
  ];
  
  for (const prefix of commonPrefixes) {
    const testPath = prefix + targetPath;
    imageFile = await zip.file(testPath);
    if (imageFile) {
      return { file: imageFile, path: testPath };
    }
  }

  // 尝试6: 不区分大小写搜索整个ZIP
  // 获取ZIP中所有文件名
  const allFiles: { name: string; file: any }[] = [];
  zip.forEach((relativePath, file) => {
    allFiles.push({ name: relativePath, file });
  });

  // 提取目标文件名（不包含路径）
  const targetFileName = targetPath.split('/').pop() || '';
  
  if (targetFileName) {
    // 搜索包含目标文件名的文件（不区分大小写）
    const lowerTarget = targetFileName.toLowerCase();
    for (const { name, file } of allFiles) {
        const fileName = name.split('/').pop() || '';
        if (fileName.toLowerCase() === lowerTarget) {
          // 确保文件在media目录中
          if (name.includes('media/')) {
            return { file, path: name };
          }
        }
    }

    // 如果没找到media目录中的文件，返回任何匹配文件名的文件
    for (const { name, file } of allFiles) {
        const fileName = name.split('/').pop() || '';
        if (fileName.toLowerCase() === lowerTarget) {
          return { file, path: name };
        }
    }
  }

  // 尝试7: 搜索任何包含"image"和正确扩展名的文件
  const imageExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'];
  for (const { name, file } of allFiles) {
    const lowerName = name.toLowerCase();
    for (const ext of imageExtensions) {
      if (lowerName.endsWith(ext)) {
        // 如果文件名包含"image"或类似模式
        if (lowerName.includes('image') || lowerName.includes('img')) {
          return { file, path: name };
        }
      }
    }
  }

  // 记录所有文件以供调试
  log('info', `Could not find image file for target: ${targetPath}`);
  log('info', `Available files in ZIP (${allFiles.length} total):`);
  const mediaFiles = allFiles.filter(({ name }) => name.includes('media/'));
  if (mediaFiles.length > 0) {
    log('info', `Media files (${mediaFiles.length}):`);
    mediaFiles.slice(0, 20).forEach(({ name }) => {
      log('info', `  - ${name}`);
    });
    if (mediaFiles.length > 20) {
      log('info', `  ... and ${mediaFiles.length - 20} more`);
    }
  }

  return null;
}

/**
 * 解析 relId 并返回 base64 URL（带缓存）
 */
async function resolveImageRelId(
  zip: JSZip,
  relId: string,
  relsMap: RelsMap,
  imageMap: Map<string, string>
): Promise<string | null> {
  // 检查缓存
  if (imageMap.has(relId)) {
    log('info', `Cache hit for relId=${relId}`);
    return imageMap.get(relId)!;
  }

  // 从 relsMap 获取目标路径
  const rel = relsMap[relId];
  if (!rel) {
    log('warn', `No relationship found for relId=${relId} in relsMap`);
    return null;
  }

  log('info', `Resolving image relId=${relId}, target=${rel.target}`);

  // 使用增强的文件查找逻辑
  const result = await findImageFileInZip(zip, rel.target);
  if (!result) {
    log('warn', `Image file not found in zip after exhaustive search. Original target: ${rel.target}`);
    return null;
  }

  const { file: imageFile, path: imagePath } = result;
  log('info', `Found image file at: ${imagePath}`);

  // 转换为 base64 URL
  const base64 = await imageFile.async('base64');
  const mimeType = inferMimeType(imagePath);
  const imageUrl = `data:${mimeType};base64,${base64}`;

  log('info', `Converted image to base64 URL, mimeType=${mimeType}, length=${imageUrl.length}`);

  // 缓存并返回
  imageMap.set(relId, imageUrl);
  return imageUrl;
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
