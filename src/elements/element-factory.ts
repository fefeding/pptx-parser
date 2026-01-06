/**
 * 从解析后的数据创建元素实例的辅助函数
 * 避免循环依赖问题
 */

import { BaseElement } from './BaseElement';
import { ShapeElement } from './ShapeElement';
import { ImageElement } from './ImageElement';
import { OleElement } from './OleElement';
import { ChartElement } from './ChartElement';
import { TableElement } from './TableElement';
import { DiagramElement } from './DiagramElement';
import { GroupElement } from './GroupElement';

/**
 * 从解析后的数据创建元素实例（用于toHTML渲染）
 * @param data 元素数据对象
 * @param relsMap 关联关系映射表
 * @param mediaMap 媒体资源映射表（relId -> base64 URL）
 * @returns BaseElement 实例或 null
 */
export function createElementFromData(
  data: any,
  relsMap: Record<string, any> = {},
  mediaMap?: Map<string, string>
): BaseElement | null {
  if (!data || !data.type) return null;

  switch (data.type) {
    case 'shape':
    case 'text': {
      const element = new ShapeElement(data.id, data.type, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'image': {
      // 如果有 mediaMap 且 relId 存在,优先使用 base64 URL
      let src = data.src || '';
      if (mediaMap && data.relId && mediaMap.has(data.relId)) {
        src = mediaMap.get(data.relId) || src;
      }

      const element = new ImageElement(data.id, data.rect, src, data.relId || '', data.props || {}, relsMap);
      Object.assign(element, data);
      // 确保 src 使用正确的值
      element.src = src;
      return element;
    }
    case 'ole': {
      const element = new OleElement(data.id, data.rect, data.progId, data.relId, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'chart': {
      const element = new ChartElement(data.id, 'chart', data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'table': {
      const element = new TableElement(data.id, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'diagram': {
      const element = new DiagramElement(data.id, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'group': {
      const element = new GroupElement(data.id, data.rect, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    default:
      return null;
  }
}
