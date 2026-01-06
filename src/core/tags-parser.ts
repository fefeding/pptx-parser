/**
 * 标签解析器
 * 解析幻灯片的自定义标签（tags）和扩展属性
 * 对齐 PPTXjs 的标签解析能力
 */

import JSZip from 'jszip';
import { getFirstChildByTagNS, getAttrSafe, log } from '../utils';
import { NS } from '../constants';
import type { RelsMap } from './types';

/**
 * 幻灯片标签
 */
export interface SlideTag {
  name: string;
  value: string;
}

/**
 * 扩展数据
 */
export interface ExtensionData {
  uri?: string;
  data?: any;
}

/**
 * 自定义属性
 */
export interface CustomProperty {
  name: string;
  value: any;
  type?: 'string' | 'number' | 'boolean' | 'date';
}

/**
 * 标签解析结果
 */
export interface TagsResult {
  id: string;
  slideId?: string;
  tags: SlideTag[];
  extensions: ExtensionData[];
  customProperties: CustomProperty[];
  relsMap: RelsMap;
}

/**
 * 解析所有幻灯片的标签
 * @param zip JSZip对象
 * @param slideCount 幻灯片数量
 * @returns 标签结果数组
 */
export async function parseAllSlideTags(zip: JSZip, slideCount: number): Promise<TagsResult[]> {
  try {
    const tagsResults: TagsResult[] = [];

    for (let i = 1; i <= slideCount; i++) {
      const slideId = `slide${i}`;
      const tagsResult = await parseSlideTags(zip, slideId);
      tagsResults.push(tagsResult);
    }

    log('info', `Parsed tags for ${slideCount} slides`);
    return tagsResults;
  } catch (error) {
    log('error', 'Failed to parse slide tags', error);
    return [];
  }
}

/**
 * 解析单个幻灯片的标签
 * @param zip JSZip对象
 * @param slideId 幻灯片ID
 * @returns 标签解析结果
 */
export async function parseSlideTags(zip: JSZip, slideId: string): Promise<TagsResult> {
  try {
    // 读取幻灯片XML
    const slidePath = `ppt/slides/${slideId}.xml`;
    const slideXml = await zip.file(slidePath)?.async('string');

    if (!slideXml) {
      log('warn', `Slide not found: ${slidePath}`);
      return createEmptyTagsResult(slideId);
    }

    // 读取关联关系
    const relsMap = await parseSlideRels(zip, slideId);

    // 解析标签
    const tags = parseTags(slideXml);

    // 解析扩展数据
    const extensions = parseExtensions(slideXml);

    // 解析自定义属性
    const customProperties = parseCustomProperties(slideXml);

    return {
      id: `tags_${slideId}`,
      slideId,
      tags,
      extensions,
      customProperties,
      relsMap
    };
  } catch (error) {
    log('error', `Failed to parse tags for slide ${slideId}`, error);
    return createEmptyTagsResult(slideId);
  }
}

/**
 * 创建空的标签结果
 */
function createEmptyTagsResult(slideId: string): TagsResult {
  return {
    id: `tags_${slideId}`,
    slideId,
    tags: [],
    extensions: [],
    customProperties: [],
    relsMap: {}
  };
}

/**
 * 解析标签
 * @param slideXml 幻灯片XML
 * @returns 标签数组
 */
function parseTags(slideXml: string): SlideTag[] {
  const tags: SlideTag[] = [];

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(slideXml, 'application/xml');
    const root = doc.documentElement;

    // 查找 extLst 元素
    const extLst = getFirstChildByTagNS(root, 'extLst', NS.p);
    if (!extLst) return tags;

    // 查找 ext 元素
    const extNodes = Array.from(extLst.children).filter(
      child => child.tagName === 'p:ext' || child.tagName.includes(':ext')
    );

    for (const ext of extNodes) {
      const uri = ext.getAttribute('uri') || '';

      // 检查是否是标签扩展
      if (uri.includes('tags') || uri.includes('ppt/tags')) {
        const tagNodes = Array.from(ext.children).filter(
          child => child.tagName === 'p:tag' || child.tagName.includes(':tag')
        );

        for (const tagNode of tagNodes) {
          const name = getAttrSafe(tagNode, 'name', '');
          const value = getAttrSafe(tagNode, 'val', '');

          if (name) {
            tags.push({ name, value });
          }
        }
      }
    }
  } catch (error) {
    log('warn', 'Failed to parse tags', error);
  }

  return tags;
}

/**
 * 解析扩展数据
 * @param slideXml 幻灯片XML
 * @returns 扩展数据数组
 */
function parseExtensions(slideXml: string): ExtensionData[] {
  const extensions: ExtensionData[] = [];

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(slideXml, 'application/xml');
    const root = doc.documentElement;

    // 查找 extLst 元素
    const extLst = getFirstChildByTagNS(root, 'extLst', NS.p);
    if (!extLst) return extensions;

    // 查找 ext 元素
    const extNodes = Array.from(extLst.children).filter(
      child => child.tagName === 'p:ext' || child.tagName.includes(':ext')
    );

    for (const ext of extNodes) {
      const uri = ext.getAttribute('uri') || '';

      // 读取扩展数据
      const childNodes = Array.from(ext.children);
      const data = parseExtensionData(childNodes);

      extensions.push({
        uri,
        data
      });
    }
  } catch (error) {
    log('warn', 'Failed to parse extensions', error);
  }

  return extensions;
}

/**
 * 解析扩展数据
 * @param childNodes 子节点数组
 * @returns 解析后的数据对象
 */
function parseExtensionData(childNodes: Element[]): any {
  const data: any = {};

  for (const child of childNodes) {
    if (child.nodeType !== 1) continue;

    const tagName = child.localName || child.tagName.split(':').pop();

    if (child.children.length > 0) {
      // 递归解析子元素
      data[tagName] = parseExtensionData(Array.from(child.children));
    } else if (child.textContent) {
      // 解析文本内容
      const text = child.textContent.trim();
      data[tagName] = text;
    } else {
      // 空元素
      data[tagName] = null;
    }
  }

  return data;
}

/**
 * 解析自定义属性
 * @param slideXml 幻灯片XML
 * @returns 自定义属性数组
 */
function parseCustomProperties(slideXml: string): CustomProperty[] {
  const properties: CustomProperty[] = [];

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(slideXml, 'application/xml');
    const root = doc.documentElement;

    // 查找 extLst 元素
    const extLst = getFirstChildByTagNS(root, 'extLst', NS.p);
    if (!extLst) return properties;

    // 查找 ext 元素
    const extNodes = Array.from(extLst.children).filter(
      child => child.tagName === 'p:ext' || child.tagName.includes(':ext')
    );

    for (const ext of extNodes) {
      const uri = ext.getAttribute('uri') || '';

      // 检查是否是自定义属性扩展
      if (uri.includes('customProps') || uri.includes('ppt/customProperties')) {
        const propNodes = Array.from(ext.children).filter(
          child => child.tagName === 'p:custData' || child.tagName.includes(':custData')
        );

        for (const propNode of propNodes) {
          const name = getAttrSafe(propNode, 'name', '');
          const value = propNode.textContent || '';
          const type = detectPropertyType(value);

          if (name) {
            properties.push({ name, value, type });
          }
        }
      }
    }
  } catch (error) {
    log('warn', 'Failed to parse custom properties', error);
  }

  return properties;
}

/**
 * 检测属性类型
 * @param value 属性值
 * @returns 属性类型
 */
function detectPropertyType(value: string): 'string' | 'number' | 'boolean' | 'date' {
  if (!value) return 'string';

  // 检查布尔值
  if (value.toLowerCase() === 'true' || value.toLowerCase() === 'false') {
    return 'boolean';
  }

  // 检查数字
  if (!isNaN(parseFloat(value)) && isFinite(Number(value))) {
    return 'number';
  }

  // 检查日期格式
  const dateRegex = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?$/;
  if (dateRegex.test(value)) {
    return 'date';
  }

  return 'string';
}

/**
 * 解析幻灯片关联关系
 */
async function parseSlideRels(zip: JSZip, slideId: string): Promise<RelsMap> {
  try {
    const relsPath = `ppt/slides/_rels/${slideId}.xml.rels`;
    const relsXml = await zip.file(relsPath)?.async('string');

    if (!relsXml) {
      log('warn', `Slide rels not found: ${relsPath}`);
      return {};
    }

    return parseRelationshipsXml(relsXml);
  } catch (error) {
    log('warn', `Failed to parse slide rels for ${slideId}`, error);
    return {};
  }
}

/**
 * 解析关联关系XML
 */
function parseRelationshipsXml(relsXml: string): RelsMap {
  const relsMap: RelsMap = {};

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');
    const root = doc.documentElement;

    const relationships = Array.from(root.children).filter(
      child => child.tagName === 'Relationship' || child.tagName.includes(':Relationship')
    );

    for (const rel of relationships) {
      const id = rel.getAttribute('Id') || '';
      const type = rel.getAttribute('Type') || '';
      const target = rel.getAttribute('Target') || '';

      if (id) {
        relsMap[id.replace('rId', '')] = {
          id,
          type,
          target
        };
      }
    }
  } catch (error) {
    log('warn', 'Failed to parse relationships XML', error);
  }

  return relsMap;
}

/**
 * 根据标签筛选幻灯片
 * @param tagsResults 所有标签结果
 * @param tagName 标签名
 * @param tagValue 标签值（可选）
 * @returns 匹配的幻灯片ID数组
 */
export function findSlidesByTag(
  tagsResults: TagsResult[],
  tagName: string,
  tagValue?: string
): string[] {
  const slideIds: string[] = [];

  for (const tagsResult of tagsResults) {
    if (!tagsResult.slideId) continue;

    const matched = tagsResult.tags.some(tag => {
      const nameMatch = tag.name === tagName;
      const valueMatch = tagValue ? tag.value === tagValue : true;
      return nameMatch && valueMatch;
    });

    if (matched) {
      slideIds.push(tagsResult.slideId);
    }
  }

  return slideIds;
}

/**
 * 根据自定义属性筛选幻灯片
 * @param tagsResults 所有标签结果
 * @param propName 属性名
 * @param propValue 属性值（可选）
 * @returns 匹配的幻灯片ID数组
 */
export function findSlidesByProperty(
  tagsResults: TagsResult[],
  propName: string,
  propValue?: any
): string[] {
  const slideIds: string[] = [];

  for (const tagsResult of tagsResults) {
    if (!tagsResult.slideId) continue;

    const matched = tagsResult.customProperties.some(prop => {
      const nameMatch = prop.name === propName;
      const valueMatch = propValue ? prop.value === propValue : true;
      return nameMatch && valueMatch;
    });

    if (matched) {
      slideIds.push(tagsResult.slideId);
    }
  }

  return slideIds;
}
