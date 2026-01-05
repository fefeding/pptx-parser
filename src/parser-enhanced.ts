/**
 * PPTX增强解析器
 * 基于原库的 parsePptx 增量扩展，完全兼容原有API
 *
 * 新增功能：
 * 1. 完整解析幻灯片元素（文本框、形状、图片、OLE对象等）
 * 2. 解析关联关系文件（rels）
 * 3. 解析元数据（作者、创建时间等）
 * 4. 图片资源解析（Base64）
 * 5. 完善的错误处理和日志
 */

import JSZip from 'jszip';
import { NS, PATHS } from './constants';
import {
  parseSlide,
  type SlideParseResult,
  type RelsMap
} from './parseSlide';
import {
  parseRels,
  parseMetadata,
  parseSlideSize,
  generateId,
  log,
  getFirstChildByTagNS
} from './utils';
import type { PptxParseResult, ParseOptions } from './types-enhanced';

/**
 * 增强版 PPTX 解析函数
 * 完全兼容原库 API，入参和出参结构保持一致
 * @param file PPTX文件（File | Blob | ArrayBuffer）
 * @param options 解析选项（可选）
 * @returns 解析结果对象
 */
export async function parsePptxEnhanced(
  file: File | Blob | ArrayBuffer,
  options?: ParseOptions
): Promise<PptxParseResult> {
  const opts = {
    parseImages: true,
    keepRawXml: false,
    verbose: false,
    ...options
  };

  try {
    log('info', 'Starting PPTX parsing...');

    // 解压ZIP文件
    const zip = await JSZip.loadAsync(file);

    // 解析元数据（docProps/core.xml）
    const metadata = await parseCoreProperties(zip);

    // 解析幻灯片尺寸
    const slideSize = await parseSlideLayoutSize(zip);

    // 解析所有幻灯片
    const slides = await parseAllSlides(zip, opts);

    // 解析全局关联关系
    const globalRels = await parseGlobalRels(zip);

    log('info', `Parsed ${slides.length} slides successfully`);

    // 计算页面比例
    const ratio = slideSize.width / slideSize.height;
    const pageSize = inferPageSize(ratio);

    const result: PptxParseResult = {
      id: generateId('ppt-doc'),
      title: metadata.title || '未命名PPT',
      author: metadata.author,
      subject: metadata.subject,
      keywords: metadata.keywords,
      description: metadata.description,
      created: metadata.created,
      modified: metadata.modified,
      slides,
      props: {
        width: slideSize.width,
        height: slideSize.height,
        ratio,
        pageSize
      },
      globalRelsMap: globalRels
    };

    // 解析图片（如果需要）
    if (opts.parseImages) {
      await parseImages(zip, result);
    }

    return result;
  } catch (error) {
    log('error', 'PPTX parsing failed', error);
    throw new Error(`PPTX解析失败: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/**
 * 解析核心属性（docProps/core.xml）
 * @param zip JSZip对象
 * @returns 元数据对象
 */
async function parseCoreProperties(zip: JSZip): Promise<{
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  created?: string;
  modified?: string;
}> {
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
 * @returns 尺寸对象
 */
async function parseSlideLayoutSize(zip: JSZip): Promise<{ width: number; height: number }> {
  try {
    // 尝试从 presentation.xml 解析
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (presentationXml) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(presentationXml, 'application/xml');

      const sldSz = doc.getElementsByTagNameNS(NS.p, 'sldSz')[0];
      if (sldSz) {
        const cx = sldSz.getAttribute('cx');
        const cy = sldSz.getAttribute('cy');

        if (cx && cy) {
          return {
            width: Math.round(parseInt(cx, 10) * 96 / 914400),
            height: Math.round(parseInt(cy, 10) * 96 / 914400)
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
 * 解析所有幻灯片
 * @param zip JSZip对象
 * @param options 解析选项
 * @returns 幻灯片数组
 */
async function parseAllSlides(
  zip: JSZip,
  options: ParseOptions
): Promise<SlideParseResult[]> {
  try {
    // 获取所有幻灯片文件
    const slideFiles = Object.keys(zip.files)
      .filter(path => path.startsWith(PATHS.SLIDES))
      .filter(path => path.endsWith('.xml'))
      .filter(path => !path.includes('_rels')) // 排除rels文件
      .sort((a, b) => {
        // 按文件名数字排序
        const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
        return numA - numB;
      });

    log('info', `Found ${slideFiles.length} slide files`);

    const slides: SlideParseResult[] = [];

    for (let i = 0; i < slideFiles.length; i++) {
      const slidePath = slideFiles[i];
      log('info', `Parsing slide ${i + 1}: ${slidePath}`);

      // 读取幻灯片XML
      const slideXml = await zip.file(slidePath)?.async('string');
      if (!slideXml) {
        log('warn', `Failed to read slide: ${slidePath}`);
        continue;
      }

      // 读取幻灯片的关联关系文件
      const slideNumber = slidePath.match(/slide(\d+)\.xml/)?.[1];
      let relsMap: RelsMap = {};

      if (slideNumber) {
        const relsPath = `${PATHS.SLIDE_RELS}slide${slideNumber}.xml.rels`;
        const relsXml = await zip.file(relsPath)?.async('string');

        if (relsXml) {
          relsMap = parseRels(relsXml);
          log('info', `Loaded ${Object.keys(relsMap).length} relationships for slide ${slideNumber}`);
        }
      }

      // 解析幻灯片
      const slide = parseSlide(slideXml, relsMap, i);

      // 保存原始XML（如果需要）
      if (options.keepRawXml) {
        slide.rawXml = slideXml;
      }

      slides.push(slide);
    }

    return slides;
  } catch (error) {
    log('error', 'Failed to parse slides', error);
    return [];
  }
}

/**
 * 解析全局关联关系
 * @param zip JSZip对象
 * @returns 关联关系映射表
 */
async function parseGlobalRels(zip: JSZip): Promise<RelsMap> {
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
 * 解析图片资源
 * @param zip JSZip对象
 * @param result 解析结果对象
 */
async function parseImages(zip: JSZip, result: PptxParseResult): Promise<void> {
  try {
    const imageMap = new Map<string, string>();

    // 遍历所有幻灯片
    for (const slide of result.slides) {
      // 遍历所有元素
      for (const element of slide.elements) {
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
 * 推断页面比例类型
 * @param ratio 宽高比
 * @returns 页面类型
 */
function inferPageSize(ratio: number): '4:3' | '16:9' | '16:10' | 'custom' {
  const epsilon = 0.01;

  if (Math.abs(ratio - 1.33333) < epsilon) return '4:3';
  if (Math.abs(ratio - 1.77778) < epsilon) return '16:9';
  if (Math.abs(ratio - 1.6) < epsilon) return '16:10';
  return 'custom';
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

// 导出主解析函数，保持向后兼容
export { parsePptxEnhanced as parsePptx };
export { parseSlide };
export type { PptxParseResult, SlideParseResult, ParseOptions };
